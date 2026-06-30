import './style.css';
import { PdfRenderer, vniToUnicode, isVniEncoding } from './src/pdf-renderer.js';
import { PdfSaver } from './src/pdf-saver.js';
import { AnnotationManager } from './src/annotation-manager.js';
import { showToast, showLoading, hideLoading } from './src/ui-utils.js';
import { PDFDocument, degrees } from 'pdf-lib';
import { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel } from 'docx';
import { saveAs } from 'file-saver';

const renderer = new PdfRenderer();
const annMgr = new AnnotationManager();

let currentFile = null;
let currentTool = 'select';
let currentColor = '#e53e3e';
let lastSelectedColor = '#e53e3e';
let fontSize = 14;
let currentFontFamily = 'Arial, Helvetica, sans-serif';
let isBold = false;
let isItalic = false;
let strokeWidth = 3;
let selectedAnn = null;
let selectedAnnEl = null;
let selectedPageIdx = null;
let drawCanvases = {};
let drawingStates = {};
let pageCount = 0;
let zoomLevel = 100;
let originalTextEdits = {};
let pageTextItems = {};
let activeTextEditor = null;

const $ = id => document.getElementById(id);
const uploadScreen = $('upload-screen');
const editorScreen = $('editor-screen');
const dropZone = $('drop-zone');
const fileInput = $('file-input');
const imageInput = $('image-input');
const pagesWrapper = $('pages-wrapper');
const thumbsContainer = $('page-thumbnails');
const fileNameDisplay = $('file-name-display');
const headerActions = $('header-actions');
const textOptions = $('text-options');
const drawOptions = $('draw-options');
const btnDelete = $('btn-delete');

// ===== EDITOR MODE (docx or scan) =====
let editorMode = 'docx'; // 'docx' = normal PDF, 'scan' = AI scan/image PDF

// ===== TAB SWITCHING =====
const tabDocx = $('tab-docx');
const tabScan = $('tab-scan');
const tabContentDocx = $('tab-content-docx');
const tabContentScan = $('tab-content-scan');

document.querySelectorAll('.upload-tab').forEach(tab => {
  tab.addEventListener('click', () => {
    document.querySelectorAll('.upload-tab').forEach(t => t.classList.remove('active'));
    document.querySelectorAll('.upload-tab-content').forEach(c => c.classList.remove('active'));
    tab.classList.add('active');
    const tabId = tab.dataset.tab;
    $('tab-content-' + tabId).classList.add('active');
  });
});

// ===== SCAN TAB API CONFIG =====
const scanApiKeyInput = $('scan-api-key');
const scanModelSelect = $('scan-model-select');
const scanCustomModelGroup = $('scan-custom-model-group');
const scanCustomModelInput = $('scan-custom-model');
const apiStatusEl = $('api-status');
const btnToggleApiKey = $('btn-toggle-api-key');

function updateScanApiStatus() {
  const key = localStorage.getItem('gemini_api_key') || '';
  if (key) {
    apiStatusEl.className = 'api-status configured';
    apiStatusEl.innerHTML = '<i class="fas fa-circle"></i><span>API Key đã được cấu hình ✓</span>';
  } else {
    apiStatusEl.className = 'api-status';
    apiStatusEl.innerHTML = '<i class="fas fa-circle"></i><span>Chưa cấu hình API Key</span>';
  }
}

// Load saved settings into scan tab fields
function loadScanApiSettings() {
  const apiKey = localStorage.getItem('gemini_api_key') || '';
  const model = localStorage.getItem('gemini_model') || 'gemini-3-flash-preview';
  const customModel = localStorage.getItem('gemini_custom_model') || '';
  scanApiKeyInput.value = apiKey;
  scanModelSelect.value = model === 'custom' ? 'custom' : 'gemini-3-flash-preview';
  // Check if saved model matches predefined options
  if (model !== 'gemini-3-flash-preview' && model !== 'custom') {
    scanModelSelect.value = 'custom';
    scanCustomModelInput.value = model;
    scanCustomModelGroup.style.display = 'block';
  } else {
    scanCustomModelInput.value = customModel;
    scanCustomModelGroup.style.display = model === 'custom' ? 'block' : 'none';
  }
  updateScanApiStatus();
}

loadScanApiSettings();

scanModelSelect.addEventListener('change', () => {
  scanCustomModelGroup.style.display = scanModelSelect.value === 'custom' ? 'block' : 'none';
});

btnToggleApiKey.addEventListener('click', () => {
  const isPassword = scanApiKeyInput.type === 'password';
  scanApiKeyInput.type = isPassword ? 'text' : 'password';
  btnToggleApiKey.querySelector('i').className = isPassword ? 'fas fa-eye-slash' : 'fas fa-eye';
});

$('btn-save-scan-api').addEventListener('click', () => {
  const apiKey = scanApiKeyInput.value.trim();
  const model = scanModelSelect.value;
  const customModel = scanCustomModelInput.value.trim();
  localStorage.setItem('gemini_api_key', apiKey);
  localStorage.setItem('gemini_model', model);
  localStorage.setItem('gemini_custom_model', customModel);
  updateScanApiStatus();
  showToast('Đã lưu cấu hình API!', 'success');
});

// ===== IMAGE TO PDF CONVERSION FOR SCAN TAB =====
async function handleScanUpload(file) {
  if (file.type.startsWith('image/')) {
    try {
      showLoading('Đang chuyển ảnh sang PDF...');
      const imgBytes = await file.arrayBuffer();
      const pdfDoc = await PDFDocument.create();
      
      let img;
      if (file.type === 'image/png') {
        img = await pdfDoc.embedPng(imgBytes);
      } else {
        // Fallback for all other images (JPEG, WebP, etc.) using canvas if needed, but embedJpg works for standard JPEGs.
        // To be safe and support all browser-supported image formats, we'll draw it to a canvas and export as PNG.
        const canvas = document.createElement('canvas');
        const imgEl = new Image();
        const objUrl = URL.createObjectURL(file);
        
        await new Promise((resolve, reject) => {
          imgEl.onload = resolve;
          imgEl.onerror = reject;
          imgEl.src = objUrl;
        });
        
        canvas.width = imgEl.width;
        canvas.height = imgEl.height;
        const ctx = canvas.getContext('2d');
        ctx.drawImage(imgEl, 0, 0);
        
        const pngDataUrl = canvas.toDataURL('image/png');
        URL.revokeObjectURL(objUrl);
        
        const res = await fetch(pngDataUrl);
        const pngBytes = await res.arrayBuffer();
        img = await pdfDoc.embedPng(pngBytes);
      }
      
      const { width, height } = img.scale(1);
      const page = pdfDoc.addPage([width, height]);
      page.drawImage(img, { x: 0, y: 0, width, height });
      
      const pdfBytes = await pdfDoc.save();
      const newName = file.name.replace(/\.[^/.]+$/, "") + ".pdf";
      const pdfFile = new File([pdfBytes], newName, { type: "application/pdf" });
      
      loadPdfScan(pdfFile);
    } catch (e) {
      console.error(e);
      showToast('Lỗi khi xử lý ảnh: ' + e.message, 'error');
      hideLoading();
    }
  } else if (file.type === 'application/pdf') {
    loadPdfScan(file);
  } else {
    showToast('Vui lòng chọn file PDF hoặc Ảnh', 'error');
  }
}

// ===== SCAN DROP ZONE =====
const dropZoneScan = $('drop-zone-scan');
const fileInputScan = $('file-input-scan');

dropZoneScan.addEventListener('click', e => {
  if (e.target.closest('.btn-upload') || e.target.closest('input')) return;
  fileInputScan.click();
});
dropZoneScan.addEventListener('dragover', e => { e.preventDefault(); dropZoneScan.classList.add('drag-over'); });
dropZoneScan.addEventListener('dragleave', () => dropZoneScan.classList.remove('drag-over'));
dropZoneScan.addEventListener('drop', e => {
  e.preventDefault();
  dropZoneScan.classList.remove('drag-over');
  const f = e.dataTransfer.files[0];
  if (f) handleScanUpload(f);
});
fileInputScan.addEventListener('change', e => { if (e.target.files[0]) handleScanUpload(e.target.files[0]); });

// ===== ZOOM CONTROLS =====
let currentZoom = 1;
function updateZoom() {
  $('zoom-level').textContent = Math.round(currentZoom * 100) + '%';
  pagesWrapper.style.zoom = currentZoom;
}

$('btn-zoom-in').addEventListener('click', () => {
  currentZoom = Math.min(currentZoom + 0.1, 3);
  updateZoom();
});

$('btn-zoom-out').addEventListener('click', () => {
  currentZoom = Math.max(currentZoom - 0.1, 0.3);
  updateZoom();
});

$('btn-zoom-fit').addEventListener('click', () => {
  currentZoom = 1;
  updateZoom();
});

// ===== ROTATE PAGE =====
$('btn-rotate-scan').addEventListener('click', async () => {
  if (!currentFile || selectedPageIdx === null) {
    // If no page is specifically clicked, rotate the first page
    if (pageCount > 0) {
      selectedPageIdx = 0;
    } else {
      showToast('Vui lòng chọn trang để xoay', 'warning');
      return;
    }
  }
  
  showLoading('Đang xoay trang...');
  try {
    const buf = await currentFile.arrayBuffer();
    const pdfDoc = await PDFDocument.load(buf);
    const pages = pdfDoc.getPages();
    
    if (selectedPageIdx >= 0 && selectedPageIdx < pages.length) {
      const page = pages[selectedPageIdx];
      const currentRot = page.getRotation().angle;
      page.setRotation(degrees((currentRot + 90) % 360));
      
      const pdfBytes = await pdfDoc.save();
      const newFile = new File([pdfBytes], currentFile.name, { type: "application/pdf" });
      
      // Reload the entire file to reflect changes and re-run OCR if in scan mode
      if (editorMode === 'scan') {
        loadPdfScan(newFile);
      } else {
        loadPdf(newFile);
      }
    }
  } catch (err) {
    console.error(err);
    showToast('Lỗi khi xoay ảnh: ' + err.message, 'error');
    hideLoading();
  }
});

// ===== AI OCR: Load scan PDF and process with Gemini =====
async function loadPdfScan(file) {
  const apiKey = localStorage.getItem('gemini_api_key') || '';
  if (!apiKey) {
    showToast('Vui lòng cấu hình API Key trước khi sử dụng tính năng AI!', 'error');
    return;
  }

  editorMode = 'scan';
  currentFile = file;
  showLoading('Đang tải PDF scan...');

  try {
    pageCount = await renderer.load(file);
    $('current-file-name').textContent = file.name;
    fileNameDisplay.style.display = 'flex';
    headerActions.style.display = 'flex';
    $('btn-save').style.display = 'none'; // Hide Save PDF button in scan mode
    uploadScreen.style.display = 'none';
    editorScreen.style.display = 'flex';
    pagesWrapper.innerHTML = '';
    thumbsContainer.innerHTML = '';
    drawCanvases = {};
    drawingStates = {};
    originalTextEdits = {};
    pageTextItems = {};

    // Show AI mode badge in toolbar
    $('toolbar-ai-indicator').style.display = 'flex';

    // Render all pages first (visual)
    for (let i = 1; i <= pageCount; i++) {
      await renderPageInEditor(i);
      await renderThumbnail(i);
    }
    hideLoading();

    // In scan mode, we no longer run automatic OCR to overlay text.
    // The user will preview the document, rotate if needed, and click "Xuất DOCX".
  } catch (e) {
    console.error(e);
    showToast('Lỗi tải PDF: ' + e.message, 'error');
  }
  hideLoading();
}

// ===== EXPORT TO DOCX (SCAN TAB) =====
$('btn-export-docx').addEventListener('click', async () => {
  if (!currentFile || pageCount === 0) return;
  const apiKey = localStorage.getItem('gemini_api_key') || '';
  if (!apiKey) {
    showToast('Vui lòng cấu hình API Key trước khi xuất DOCX!', 'error');
    return;
  }

  showLoading('Đang phân tích cấu trúc và nhận diện văn bản...');
  let selectedModel = localStorage.getItem('gemini_model') || 'gemini-3-flash-preview';
  if (selectedModel === 'custom') {
    selectedModel = localStorage.getItem('gemini_custom_model') || 'gemini-3-flash-preview';
  }

  try {
    const allDocxElements = [];
    let processedPages = 0;

    for (let i = 0; i < pageCount; i++) {
      const progressPct = Math.round(((i) / pageCount) * 100);
      $('loading-text').textContent = `AI đang xử lý trang ${i + 1}/${pageCount} (${progressPct}%)...`;
      
      const wrapper = pagesWrapper.children[i];
      const pdfCanvas = wrapper.querySelector('.pdf-render');
      const pageDataUrl = pdfCanvas.toDataURL('image/png');

      const promptText = `You are an expert document layout analyzer.
Extract all content from this document page image and represent it as a JSON array of formatting blocks.
Preserve the original reading order, structure, headings, paragraphs, and lists.

For each block, output an object:
1. "type": "heading" or "paragraph" or "list"
2. "level": (only for heading) 1, 2, or 3
3. "alignment": "left", "center", "right", or "justify"
4. "runs": An array of text segments, each with:
    - "text": The text content
    - "bold": true/false
    - "italic": true/false

Example output:
[
  { "type": "heading", "level": 1, "alignment": "center", "runs": [{ "text": "TÊN CƠ QUAN", "bold": true, "italic": false }] },
  { "type": "paragraph", "alignment": "left", "runs": [{ "text": "Hôm nay là ngày...", "bold": false, "italic": false }, { "text": " rất quan trọng", "bold": true, "italic": false }] }
]

Return STRICTLY valid JSON only. Do not include markdown code block syntax (like \`\`\`json) or any explanations.`;

      const responseText = await callGeminiVisionAPI(apiKey, selectedModel, pageDataUrl, promptText);

      let blocks = [];
      try {
        let cleanText = responseText.trim();
        if (cleanText.startsWith('\`\`\`')) {
          cleanText = cleanText.replace(/^\`\`\`json\s*/, '').replace(/\`\`\`$/, '').trim();
        }
        blocks = JSON.parse(cleanText);
      } catch (err) {
        console.error("Failed to parse Gemini OCR JSON:", responseText, err);
        throw new Error("Không thể phân tích định dạng từ phản hồi AI.");
      }

      for (const block of blocks) {
        const textRuns = (block.runs || []).map(r => new TextRun({
          text: r.text || "",
          bold: r.bold || false,
          italics: r.italic || false,
          font: "Times New Roman"
        }));

        let alignment = AlignmentType.LEFT;
        if (block.alignment === 'center') alignment = AlignmentType.CENTER;
        if (block.alignment === 'right') alignment = AlignmentType.RIGHT;
        if (block.alignment === 'justify') alignment = AlignmentType.JUSTIFIED;

        if (block.type === 'heading') {
          let hLevel = HeadingLevel.HEADING_1;
          if (block.level === 2) hLevel = HeadingLevel.HEADING_2;
          if (block.level === 3) hLevel = HeadingLevel.HEADING_3;
          allDocxElements.push(new Paragraph({ children: textRuns, alignment, heading: hLevel, spacing: { before: 200, after: 120 } }));
        } else if (block.type === 'list') {
          allDocxElements.push(new Paragraph({ children: textRuns, alignment, bullet: { level: 0 }, spacing: { after: 120 } }));
        } else {
          // Paragraph fallback
          allDocxElements.push(new Paragraph({ children: textRuns, alignment, spacing: { after: 120 } }));
        }
      }
      processedPages++;
    }

    const doc = new Document({
      sections: [{
        properties: {},
        children: allDocxElements
      }]
    });

    const blob = await Packer.toBlob(doc);
    saveAs(blob, currentFile.name.replace(/\.[^/.]+$/, "") + ".docx");
    
    showToast(`Xuất file Word thành công (${processedPages} trang)!`, 'success');
  } catch (err) {
    console.error(err);
    showToast('Lỗi xuất DOCX: ' + err.message, 'error');
  }
  hideLoading();
});

// ===== FILE UPLOAD (DOCX mode) =====
dropZone.addEventListener('click', e => {
  if (e.target.closest('.btn-upload') || e.target.closest('input')) return;
  fileInput.click();
});
dropZone.addEventListener('dragover', e => { e.preventDefault(); dropZone.classList.add('drag-over'); });
dropZone.addEventListener('dragleave', () => dropZone.classList.remove('drag-over'));
dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const f = e.dataTransfer.files[0];
  if (f && f.type === 'application/pdf') loadPdf(f);
  else showToast('Vui lòng chọn file PDF', 'error');
});
fileInput.addEventListener('change', e => { if (e.target.files[0]) loadPdf(e.target.files[0]); });

// ===== LOAD PDF =====
async function loadPdf(file) {
  editorMode = 'docx';
  currentFile = file;
  showLoading('Đang tải PDF...');
  try {
    // Hide AI mode badge in toolbar for normal mode
    $('toolbar-ai-indicator').style.display = 'none';
    pageCount = await renderer.load(file);
    $('current-file-name').textContent = file.name;
    fileNameDisplay.style.display = 'flex';
    headerActions.style.display = 'flex';
    $('btn-save').style.display = 'inline-flex'; // Show Save PDF button
    uploadScreen.style.display = 'none';
    editorScreen.style.display = 'flex';
    pagesWrapper.innerHTML = '';
    thumbsContainer.innerHTML = '';
    drawCanvases = {};
    drawingStates = {};
    originalTextEdits = {};
    pageTextItems = {};

    for (let i = 1; i <= pageCount; i++) {
      await renderPageInEditor(i);
      await renderThumbnail(i);
    }
    showToast(`Đã tải ${pageCount} trang — Click vào chữ để sửa trực tiếp`, 'success');
  } catch (e) {
    console.error(e);
    showToast('Lỗi tải PDF: ' + e.message, 'error');
  }
  hideLoading();
}

async function renderPageInEditor(num) {
  const idx = num - 1;
  const wrapper = document.createElement('div');
  wrapper.className = 'pdf-page-wrapper';
  wrapper.dataset.page = idx;

  const pdfCanvas = document.createElement('canvas');
  pdfCanvas.className = 'pdf-render';
  wrapper.appendChild(pdfCanvas);
  const { width, height } = await renderer.renderPage(num, pdfCanvas);

  // Wait for all custom web fonts from the PDF to be fully ready in the document
  if (document.fonts && document.fonts.ready) {
    await document.fonts.ready;
  }

  // Drawing overlay canvas
  const drawCanvas = document.createElement('canvas');
  drawCanvas.className = 'annotation-layer';
  drawCanvas.width = width;
  drawCanvas.height = height;
  wrapper.appendChild(drawCanvas);
  drawCanvases[idx] = drawCanvas;
  drawingStates[idx] = { paths: [], currentPath: null };
  setupDrawCanvas(drawCanvas, idx);

  // Editable text layer
  const textLayer = document.createElement('div');
  textLayer.className = 'text-annotations';
  wrapper.appendChild(textLayer);

  // Extract original text and make each piece directly editable
  try {
    const textItems = await renderer.getTextContent(num);
    pageTextItems[idx] = textItems;
    textItems.forEach(item => {
      createInlineEditor(item, idx, textLayer, width, height, pdfCanvas);
    });
  } catch (e) {
    console.warn('Text extraction failed page', num, e);
  }

  // Click on empty area = add new text (only in select/text mode)
  wrapper.addEventListener('dblclick', e => {
    if (e.target.closest('.pdf-text-overlay') || e.target.closest('.text-annotation') || e.target.closest('.image-annotation')) return;
    if (currentTool === 'draw' || currentTool === 'highlight' || currentTool === 'rect' || currentTool === 'eraser' || currentTool === 'cut' || currentTool === 'erase-stamp') return;
    const rect = textLayer.getBoundingClientRect();
    const scaleX = width / rect.width;
    const scaleY = height / rect.height;
    const x = (e.clientX - rect.left) * scaleX;
    const y = (e.clientY - rect.top) * scaleY;
    createNewTextAnnotation(idx, textLayer, x, y, width, height);
  });

  pagesWrapper.appendChild(wrapper);
}

// ===== COLOR DETECTION UTILITY FOR CANVAS PIXELS =====
function detectColors(canvas, x, y, w, h) {
  const ctx = canvas.getContext('2d');
  
  // Make sure bounds are integer values and inside canvas dimensions
  const rx = Math.max(0, Math.floor(x));
  const ry = Math.max(0, Math.floor(y));
  const rw = Math.min(canvas.width - rx, Math.ceil(w));
  const rh = Math.min(canvas.height - ry, Math.ceil(h));
  
  if (rw <= 0 || rh <= 0) {
    return { text: 'rgb(0,0,0)', bg: 'rgb(255,255,255)', bgRgb: { r: 1, g: 1, b: 1 } };
  }
  
  try {
    const imgData = ctx.getImageData(rx, ry, rw, rh);
    const data = imgData.data;
    
    // Count color frequency to find background color (most common color in the bounding box)
    const counts = {};
    let maxCount = 0;
    let bgColor = 'rgb(255,255,255)';
    let bgR = 255, bgG = 255, bgB = 255;
    
    for (let i = 0; i < data.length; i += 4) {
      const r = data[i];
      const g = data[i+1];
      const b = data[i+2];
      const a = data[i+3];
      if (a < 50) continue; // Skip semi-transparent pixels
      
      const rgbStr = `rgb(${r},${g},${b})`;
      counts[rgbStr] = (counts[rgbStr] || 0) + 1;
      if (counts[rgbStr] > maxCount) {
        maxCount = counts[rgbStr];
        bgColor = rgbStr;
        bgR = r;
        bgG = g;
        bgB = b;
      }
    }
    
    const bgBrightness = (bgR * 299 + bgG * 587 + bgB * 114) / 1000;
    
    // Find contrasting pixels to determine text color
    const contrastingPixels = [];
    for (let i = 0; i < data.length; i += 4) {
      const r = data[i];
      const g = data[i+1];
      const b = data[i+2];
      const a = data[i+3];
      if (a < 50) continue;
      
      const dist = Math.abs(r - bgR) + Math.abs(g - bgG) + Math.abs(b - bgB);
      if (dist > 60) { // Threshold for contrasting from background
        const brightness = (r * 299 + g * 587 + b * 114) / 1000;
        contrastingPixels.push({ r, g, b, brightness });
      }
    }
    
    let textColor = 'rgb(0,0,0)';
    if (contrastingPixels.length > 0) {
      // For light background, sort ascending (find darkest pixels representing true font color)
      // For dark background, sort descending (find lightest pixels representing true font color)
      if (bgBrightness > 128) {
        contrastingPixels.sort((a, b) => a.brightness - b.brightness);
      } else {
        contrastingPixels.sort((a, b) => b.brightness - a.brightness);
      }
      
      // Average the top 15% extreme contrasting pixels to filter out anti-aliasing edges
      const sampleSize = Math.max(1, Math.floor(contrastingPixels.length * 0.15));
      let sumR = 0, sumG = 0, sumB = 0;
      for (let i = 0; i < sampleSize; i++) {
        sumR += contrastingPixels[i].r;
        sumG += contrastingPixels[i].g;
        sumB += contrastingPixels[i].b;
      }
      const finalR = Math.round(sumR / sampleSize);
      const finalG = Math.round(sumG / sampleSize);
      const finalB = Math.round(sumB / sampleSize);
      textColor = `rgb(${finalR},${finalG},${finalB})`;
    } else {
      textColor = bgBrightness > 128 ? 'rgb(0,0,0)' : 'rgb(255,255,255)';
    }
    
    return {
      text: textColor,
      bg: bgColor,
      bgRgb: { r: bgR / 255, g: bgG / 255, b: bgB / 255 }
    };
  } catch (e) {
    console.warn('Color detection failed, using defaults', e);
    return { text: 'rgb(0,0,0)', bg: 'rgb(255,255,255)', bgRgb: { r: 1, g: 1, b: 1 } };
  }
}

// ===== INLINE EDITOR FOR EXISTING PDF TEXT =====
function saveInlineEdit(el) {
  const item = el._origItem;
  const pageIdx = el._pageIdx;
  
  // Clean HTML
  let cleanedHTML = el.innerHTML
    .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\xA0\u2000-\u200A\u202F\u205F\u3000\uF020]/g, ' ')
    .replace(/[\xAD\u200B-\u200F\uFEFF]/g, '');

  if (isVniEncoding(el.textContent)) {
    const unicodeText = vniToUnicode(el.textContent);
    el.textContent = unicodeText;
    cleanedHTML = el.innerHTML;
  } else {
    el.innerHTML = cleanedHTML;
  }
  const newHTML = cleanedHTML;
  const newText = el.textContent;
  
  const isBold = el.style.fontWeight === 'bold';
  const isItalic = el.style.fontStyle === 'italic';
  
  if (newHTML !== el._origHTML || isBold !== el._origBold || isItalic !== el._origItalic || el.style.color !== el._origColor) {
    el.classList.add('modified');
    if (!originalTextEdits[pageIdx]) originalTextEdits[pageIdx] = [];
    originalTextEdits[pageIdx] = originalTextEdits[pageIdx].filter(e => e.origItem !== item);
    originalTextEdits[pageIdx].push({ 
      origItem: item, 
      newText, 
      newHtml: newHTML,
      isBold, 
      isItalic,
      textColor: el.style.color
    });
  } else {
    el.classList.remove('modified');
    if (originalTextEdits[pageIdx]) {
      originalTextEdits[pageIdx] = originalTextEdits[pageIdx].filter(e => e.origItem !== item);
    }
  }
}

function createInlineEditor(item, pageIdx, layer, pw, ph, pdfCanvas) {
  const el = document.createElement('div');
  el.className = 'pdf-text-overlay';
  el.style.left = (item.x / pw * 100) + '%';
  el.style.top = (item.y / ph * 100) + '%';
  
  // Set size dynamically so it grows horizontally rather than wrapping
  el.style.minWidth = Math.max(item.width + 4, 30) + 'px';
  el.style.width = 'max-content';
  el.style.height = (item.height + 4) + 'px';
  el.style.fontSize = item.fontSize + 'px';
  el.style.lineHeight = '1.35';
  el.textContent = item.str;
  el.spellcheck = false;
  el.contentEditable = true;

  // Use the resolved standard font family to prevent missing space glyph box characters
  el.style.fontFamily = item.fontFamily || 'Arial, Helvetica, sans-serif';

  // Automatically detect bold and italic styling from original PDF font name
  const fn = (item.fontName || '').toLowerCase();
  const detectedBold = fn.includes('bold') || fn.includes('gothic') || fn.includes('heavy') || fn.includes('black');
  const detectedItalic = fn.includes('italic') || fn.includes('oblique');
  
  // Since we use standard system fonts (like Arial/Helvetica) for inline editing overlays,
  // we can and should apply bold/italic styling directly so they render correctly in the browser.
  el.style.fontWeight = detectedBold ? 'bold' : 'normal';
  el.style.fontStyle = detectedItalic ? 'italic' : 'normal';

  el._origItem = item;
  el._origText = item.str;
  el._origBold = detectedBold;
  el._origItalic = detectedItalic;
  el._pageIdx = pageIdx;
  el._origHTML = el.innerHTML;

  // Perform transition to in-place edit on first focus
  el.addEventListener('focus', () => {
    if (currentTool === 'draw' || currentTool === 'highlight' || currentTool === 'rect' || currentTool === 'eraser' || currentTool === 'cut' || currentTool === 'erase-stamp') {
      el.blur();
      return;
    }
    
    // Lazy activate the block: erase the original text from the canvas and show HTML text
    if (!el.classList.contains('activated')) {
      const colors = detectColors(pdfCanvas, item.x, item.y, item.width, item.height);
      
      // Save details for export
      item.textColor = colors.text;
      item.bgColor = colors.bg;
      item.bgRgb = colors.bgRgb;
      
      // Paint over the canvas text using the detected background color
      const ctx = pdfCanvas.getContext('2d');
      ctx.fillStyle = colors.bg;
      // Clear a generously larger area to fully cover anti-aliased edges of characters
      ctx.fillRect(item.x - 3, item.y - 3, item.width + 8, item.height + 6);
      
      // Set matching text color and caret color
      el.style.color = colors.text;
      // Read back the browser-normalized color value (browser adds spaces: "rgb(0,0,0)" → "rgb(0, 0, 0)")
      // This ensures exact match when saveInlineEdit compares el.style.color vs el._origColor
      el._origColor = el.style.color;
      el.style.caretColor = colors.text;
      el.classList.add('activated');
    }
    
    el.classList.add('editing');
    activeTextEditor = el;
    el.addEventListener('keyup', saveCurrentSelection);
    el.addEventListener('mouseup', saveCurrentSelection);
    
    // Sync toolbar button states and global formatting variables
    isBold = el.style.fontWeight === 'bold';
    isItalic = el.style.fontStyle === 'italic';
    $('btn-bold').classList.toggle('active', isBold);
    $('btn-italic').classList.toggle('active', isItalic);
    selectedPageIdx = pageIdx;
    if (el.style.color) {
      $('color-picker').value = rgbToHex(el.style.color);
      currentColor = $('color-picker').value;
    }
    if (el.style.fontSize) {
      const sz = parseInt(el.style.fontSize);
      if (sz > 0) {
        $('font-size').value = sz;
        fontSize = sz;
      }
    }
    if (el.style.fontFamily) {
      syncFontFamilyDropdown(el.style.fontFamily);
    }
  });

  // Save changes on blur (but not when focus moves to toolbar)
  el.addEventListener('blur', (e) => {
    const toolbar = document.getElementById('toolbar');
    // Check relatedTarget first (immediate)
    if (e.relatedTarget && toolbar && toolbar.contains(e.relatedTarget)) {
      return; // Focus moved to toolbar — keep editor active
    }
    // Fallback: check after a tick (for native popups like color picker where relatedTarget is null)
    setTimeout(() => {
      const active = document.activeElement;
      if (active && toolbar && toolbar.contains(active)) {
        return; // Focus is on toolbar — keep editor active
      }
      el.classList.remove('editing');
      saveInlineEdit(el);
      setTimeout(() => { if (activeTextEditor === el) activeTextEditor = null; }, 200);
    }, 0);
  });

  // Hotkey behaviors
  el.addEventListener('keydown', e => {
    if (e.key === 'Escape') {
      el.textContent = el._origText;
      el.style.fontWeight = el._origBold ? 'bold' : 'normal';
      el.style.fontStyle = el._origItalic ? 'italic' : 'normal';
      el.blur();
    }
    e.stopPropagation();
  });

  layer.appendChild(el);
}

// ===== FIND NEAREST ORIGINAL TEXT ITEM =====
function findNearestTextItem(pageIdx, x, y) {
  const items = pageTextItems[pageIdx];
  if (!items || items.length === 0) return null;
  let nearestItem = null;
  let minDist = Infinity;
  for (const item of items) {
    const dx = x - (item.x + item.width / 2);
    const dy = y - (item.y + item.height / 2);
    const dist = Math.sqrt(dx * dx + dy * dy);
    if (dist < minDist) {
      minDist = dist;
      nearestItem = item;
    }
  }

  if (nearestItem && !nearestItem.textColor) {
    const wrapper = pagesWrapper.children[pageIdx];
    const pdfCanvas = wrapper ? wrapper.querySelector('.pdf-render') : null;
    if (pdfCanvas) {
      const colors = detectColors(pdfCanvas, nearestItem.x, nearestItem.y, nearestItem.width, nearestItem.height);
      nearestItem.textColor = colors.text;
      nearestItem.bgColor = colors.bg;
      nearestItem.bgRgb = colors.bgRgb;
    }
  }
  return nearestItem;
}

// ===== ADD NEW TEXT (double-click on empty area) =====
function createNewTextAnnotation(pageIdx, layer, x, y, pw, ph) {
  // New text annotations always use standard system fonts, never PDF-embedded fonts.
  // This ensures user's toolbar formatting (bold, italic, color) is always respected.
  const fontFamily = currentFontFamily;

  const ann = annMgr.add(pageIdx, {
    type: 'text', x, y, text: '', color: currentColor,
    fontSize: fontSize, bold: isBold, italic: isItalic, fontName: null, fontFamily
  });
  renderNewTextElement(ann, pageIdx, layer, pw, ph, true);
}

function createNewTextAnnotationWithStyle(pageIdx, layer, x, y, w, h, pw, ph) {
  const fontFamily = currentFontFamily;

  const ann = annMgr.add(pageIdx, {
    type: 'text', x, y, w, h, text: '', color: currentColor,
    fontSize: fontSize, bold: isBold, italic: isItalic, fontName: null, fontFamily
  });
  renderNewTextElement(ann, pageIdx, layer, pw, ph, true);
}

function renderNewTextElement(ann, pageIdx, layer, pw, ph, autoFocus = false) {
  const el = document.createElement('div');
  el.className = 'text-annotation new-text';
  el.contentEditable = true;
  el.dataset.id = ann.id;
  el.style.left = (ann.x / pw * 100) + '%';
  el.style.top = (ann.y / ph * 100) + '%';
  
  if (ann.w) {
    el.style.width = (ann.w / pw * 100) + '%';
  } else {
    el.style.minWidth = '60px';
    el.style.width = 'max-content';
  }
  if (ann.h) {
    el.style.height = (ann.h / ph * 100) + '%';
  } else {
    el.style.minHeight = '20px';
  }
  
  el.style.color = ann.color || currentColor;
  el.style.fontSize = ann.fontSize + 'px';
  el.style.fontFamily = ann.fontFamily || 'Arial, Helvetica, sans-serif';
  // New text annotations always use standard system fonts, so bold/italic always works.
  // Never suppress bold/italic based on PDF-embedded fontName for user-created annotations.
  el.style.fontWeight = ann.bold ? 'bold' : 'normal';
  el.style.fontStyle = ann.italic ? 'italic' : 'normal';
  if (ann.html) {
    el.innerHTML = ann.html;
  } else {
    el.textContent = ann.text;
  }
  el.spellcheck = false;

  // Drag support
  let dragging = false, dx, dy;
  el.addEventListener('mousedown', e => {
    if ((currentTool !== 'select' && currentTool !== 'text') || el.classList.contains('editing')) return;
    dragging = true;
    dx = e.clientX - el.getBoundingClientRect().left;
    dy = e.clientY - el.getBoundingClientRect().top;
    selectAnnotation(ann, pageIdx, el);
    e.preventDefault();
  });
  document.addEventListener('mousemove', e => {
    if (!dragging) return;
    const rect = layer.getBoundingClientRect();
    el.style.left = ((e.clientX - rect.left - dx) / rect.width * 100) + '%';
    el.style.top = ((e.clientY - rect.top - dy) / rect.height * 100) + '%';
  });
  document.addEventListener('mouseup', () => {
    if (dragging) {
      dragging = false;
      annMgr.update(pageIdx, ann.id, {
        x: parseFloat(el.style.left) / 100 * pw,
        y: parseFloat(el.style.top) / 100 * ph
      });
    }
  });

  el.addEventListener('dblclick', e => {
    e.stopPropagation();
    el.classList.add('editing');
    el.focus();
  });

  el.addEventListener('focus', () => {
    activeTextEditor = el;
    el.addEventListener('keyup', saveCurrentSelection);
    el.addEventListener('mouseup', saveCurrentSelection);
    // Sync toolbar button states and global formatting variables
    isBold = el.style.fontWeight === 'bold';
    isItalic = el.style.fontStyle === 'italic';
    $('btn-bold').classList.toggle('active', isBold);
    $('btn-italic').classList.toggle('active', isItalic);
    selectedPageIdx = pageIdx;
    if (el.style.color) {
      $('color-picker').value = rgbToHex(el.style.color);
      currentColor = $('color-picker').value;
    }
    if (el.style.fontSize) {
      const sz = parseInt(el.style.fontSize);
      if (sz > 0) {
        $('font-size').value = sz;
        fontSize = sz;
      }
    }
    if (el.style.fontFamily) {
      syncFontFamilyDropdown(el.style.fontFamily);
    }
  });

  el.addEventListener('blur', (e) => {
    const toolbar = document.getElementById('toolbar');
    // Check relatedTarget first (immediate)
    if (e.relatedTarget && toolbar && toolbar.contains(e.relatedTarget)) {
      return; // Focus moved to toolbar — keep editor active
    }
    // Fallback: check after a tick (for native popups like color picker where relatedTarget is null)
    setTimeout(() => {
      const active = document.activeElement;
      if (active && toolbar && toolbar.contains(active)) {
        return; // Focus is on toolbar — keep editor active
      }
      el.classList.remove('editing');
      let cleanedHTML = el.innerHTML
        .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\xA0\u2000-\u200A\u202F\u205F\u3000\uF020]/g, ' ')
        .replace(/[\xAD\u200B-\u200F\uFEFF]/g, '');

      if (isVniEncoding(el.textContent)) {
        const unicodeText = vniToUnicode(el.textContent);
        el.textContent = unicodeText;
        cleanedHTML = el.innerHTML;
      } else {
        el.innerHTML = cleanedHTML;
      }

      const newText = el.textContent;
      annMgr.update(pageIdx, ann.id, { text: newText, html: cleanedHTML });
      if (!newText.trim()) {
        annMgr.remove(pageIdx, ann.id);
        el.remove();
      }
      setTimeout(() => { if (activeTextEditor === el) activeTextEditor = null; }, 200);
    }, 0);
  });

  el.addEventListener('click', e => {
    if (currentTool === 'select' || currentTool === 'text') { 
      selectAnnotation(ann, pageIdx, el); 
      e.stopPropagation(); 
    }
  });

  el.addEventListener('keydown', e => e.stopPropagation());

  layer.appendChild(el);
  if (autoFocus) {
    el.classList.add('editing');
    setTimeout(() => el.focus(), 50);
  }
  return el;
}

function selectAnnotation(ann, pageIdx, el) {
  document.querySelectorAll('.text-annotation.selected, .image-annotation.selected').forEach(e => e.classList.remove('selected'));
  el.classList.add('selected');
  selectedAnn = ann;
  selectedAnnEl = el;
  selectedPageIdx = pageIdx;
  btnDelete.style.display = 'inline-flex';

  if (ann && ann.type === 'text') {
    activeTextEditor = el;
    // Sync toolbar states and global formatting variables
    isBold = el.style.fontWeight === 'bold';
    isItalic = el.style.fontStyle === 'italic';
    $('btn-bold').classList.toggle('active', isBold);
    $('btn-italic').classList.toggle('active', isItalic);
    if (el.style.color) {
      $('color-picker').value = rgbToHex(el.style.color);
      currentColor = $('color-picker').value;
    }
    if (el.style.fontSize) {
      const sz = parseInt(el.style.fontSize);
      if (sz > 0) {
        $('font-size').value = sz;
        fontSize = sz;
      }
    }
    if (el.style.fontFamily) {
      syncFontFamilyDropdown(el.style.fontFamily);
    }
  } else {
    activeTextEditor = null;
  }
}

function deleteSelected() {
  if (!selectedAnn) return;
  annMgr.remove(selectedPageIdx, selectedAnn.id);
  refreshAnnotations(selectedPageIdx);
  selectedAnn = null;
  selectedAnnEl = null;
  btnDelete.style.display = 'none';
  showToast('Đã xóa', 'info');
}

function refreshAnnotations(pageIdx) {
  const wrapper = pagesWrapper.children[pageIdx];
  if (!wrapper) return;
  const layer = wrapper.querySelector('.text-annotations');
  const pdfCanvas = wrapper.querySelector('.pdf-render');
  // Only remove user-added annotations, keep pdf-text-overlay
  layer.querySelectorAll('.text-annotation, .image-annotation').forEach(el => el.remove());
  const pw = pdfCanvas.width, ph = pdfCanvas.height;
  annMgr.getPage(pageIdx).forEach(ann => {
    if (ann.type === 'text') {
      const el = renderNewTextElement(ann, pageIdx, layer, pw, ph);
      if (selectedAnn && selectedAnn.id === ann.id) {
        el.classList.add('selected');
        selectedAnnEl = el;
      }
    } else if (ann.type === 'image') {
      const el = renderImageElement(ann, pageIdx, layer, pw, ph);
      if (selectedAnn && selectedAnn.id === ann.id) {
        el.classList.add('selected');
        selectedAnnEl = el;
      }
    }
  });
}

async function renderThumbnail(num) {
  const div = document.createElement('div');
  div.className = 'page-thumb' + (num === 1 ? ' active' : '');
  const c = document.createElement('canvas');
  div.appendChild(c);
  const label = document.createElement('span');
  label.className = 'page-thumb-label';
  label.textContent = num;
  div.appendChild(label);
  div.addEventListener('click', () => {
    document.querySelectorAll('.page-thumb').forEach(t => t.classList.remove('active'));
    div.classList.add('active');
    const target = pagesWrapper.children[num - 1];
    if (target) target.scrollIntoView({ behavior: 'smooth', block: 'center' });
  });
  thumbsContainer.appendChild(div);
  await renderer.renderThumb(num, c);
}

// ===== TOOLBAR =====
document.querySelectorAll('[data-tool]').forEach(btn => {
  btn.addEventListener('click', () => {
    const prevTool = currentTool;
    document.querySelectorAll('[data-tool]').forEach(b => b.classList.remove('active'));
    btn.classList.add('active');
    currentTool = btn.dataset.tool;

    // Auto color switch for eraser tool
    if (currentTool === 'eraser') {
      if (prevTool !== 'eraser') {
        lastSelectedColor = currentColor;
      }
      currentColor = '#ffffff';
      $('color-picker').value = '#ffffff';
    } else if (prevTool === 'eraser') {
      currentColor = lastSelectedColor;
      $('color-picker').value = lastSelectedColor;
    }

    // Show text formatting options in both 'select' and 'text' modes
    // so users can pre-set bold/italic/size before creating or editing text
    textOptions.style.display = (currentTool === 'text' || currentTool === 'select') ? 'flex' : 'none';
    drawOptions.style.display = (currentTool === 'draw' || currentTool === 'highlight') ? 'flex' : 'none';
    updateCursors();
    if (currentTool === 'image') imageInput.click();
  });
});

function updateCursors() {
  const isDrawingTool = (currentTool === 'draw' || currentTool === 'highlight' || currentTool === 'rect' || currentTool === 'eraser' || currentTool === 'cut' || currentTool === 'erase-stamp');
  if (pagesWrapper) {
    pagesWrapper.classList.toggle('drawing-mode', isDrawingTool);
  }
  document.querySelectorAll('.annotation-layer').forEach(c => {
    if (isDrawingTool || currentTool === 'text') {
      c.style.pointerEvents = 'auto';
      c.style.cursor = 'crosshair';
    } else {
      c.style.pointerEvents = 'none';
    }
  });
}

function rgbToHex(rgbStr) {
  if (!rgbStr) return '#000000';
  if (rgbStr.startsWith('#')) return rgbStr;
  const match = rgbStr.match(/\d+/g);
  if (!match || match.length < 3) return '#000000';
  const r = parseInt(match[0]);
  const g = parseInt(match[1]);
  const b = parseInt(match[2]);
  return "#" + ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1);
}

function syncFontFamilyDropdown(fontFamilyValue) {
  if (!fontFamilyValue) return;
  const cleanFont = fontFamilyValue.replace(/['"]/g, '').split(',')[0].trim().toLowerCase();
  const options = Array.from($('font-family').options);
  const foundOption = options.find(opt => {
    const valClean = opt.value.replace(/['"]/g, '').split(',')[0].trim().toLowerCase();
    const txtClean = opt.text.trim().toLowerCase();
    return valClean.includes(cleanFont) || cleanFont.includes(valClean) || txtClean.includes(cleanFont) || cleanFont.includes(txtClean);
  });
  if (foundOption) {
    $('font-family').value = foundOption.value;
    currentFontFamily = foundOption.value;
  }
}

let savedSelectionRange = null;

function saveCurrentSelection() {
  if (!activeTextEditor) return;
  const sel = window.getSelection();
  if (sel.rangeCount > 0) {
    const range = sel.getRangeAt(0);
    if (activeTextEditor.contains(range.commonAncestorContainer)) {
      savedSelectionRange = range.cloneRange();
    }
  }
}

function restoreSelection() {
  if (!activeTextEditor || !savedSelectionRange) return;
  activeTextEditor.focus();
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(savedSelectionRange);
}

function applyTextFormat(formatType, value) {
  if (!activeTextEditor) return;

  restoreSelection();

  const selection = window.getSelection();
  const hasSelection = selection.rangeCount > 0 && 
                        activeTextEditor.contains(selection.anchorNode) && 
                        !selection.isCollapsed;

  if (hasSelection) {
    if (formatType === 'bold') {
      document.execCommand('bold', false, null);
    } else if (formatType === 'italic') {
      document.execCommand('italic', false, null);
    } else if (formatType === 'color') {
      document.execCommand('foreColor', false, value);
    } else if (formatType === 'fontFamily') {
      document.execCommand('fontName', false, value);
    } else if (formatType === 'fontSize') {
      // Use the size 7 trick to apply custom px font size
      document.execCommand('fontSize', false, '7');
      const fontEls = activeTextEditor.querySelectorAll('font[size="7"]');
      fontEls.forEach(fontEl => {
        const span = document.createElement('span');
        span.style.fontSize = `${value}px`;
        while (fontEl.firstChild) {
          span.appendChild(fontEl.firstChild);
        }
        fontEl.parentNode.replaceChild(span, fontEl);
      });
    }
    
    // Save the changes back to state (original edits or annotation manager)
    const id = activeTextEditor.dataset.id;
    if (id) {
      annMgr.update(selectedPageIdx, id, { 
        text: activeTextEditor.textContent, 
        html: activeTextEditor.innerHTML 
      });
    } else if (activeTextEditor.classList.contains('pdf-text-overlay')) {
      saveInlineEdit(activeTextEditor);
    }

    // Keep the selection updated
    saveCurrentSelection();
  } else {
    // Fallback: apply to entire block
    if (formatType === 'bold') {
      const isCurrentlyBold = activeTextEditor.style.fontWeight === 'bold';
      activeTextEditor.style.fontWeight = isCurrentlyBold ? 'normal' : 'bold';
      isBold = !isCurrentlyBold;
      $('btn-bold').classList.toggle('active', isBold);
    } else if (formatType === 'italic') {
      const isCurrentlyItalic = activeTextEditor.style.fontStyle === 'italic';
      activeTextEditor.style.fontStyle = isCurrentlyItalic ? 'normal' : 'italic';
      isItalic = !isCurrentlyItalic;
      $('btn-italic').classList.toggle('active', isItalic);
    } else if (formatType === 'color') {
      activeTextEditor.style.color = value;
    } else if (formatType === 'fontFamily') {
      activeTextEditor.style.fontFamily = value;
    } else if (formatType === 'fontSize') {
      activeTextEditor.style.fontSize = `${value}px`;
    }

    const id = activeTextEditor.dataset.id;
    if (id) {
      const updateData = {
        text: activeTextEditor.textContent,
        html: activeTextEditor.innerHTML
      };
      if (formatType === 'bold') updateData.bold = isBold;
      if (formatType === 'italic') updateData.italic = isItalic;
      if (formatType === 'color') updateData.color = value;
      if (formatType === 'fontFamily') updateData.fontFamily = value;
      if (formatType === 'fontSize') updateData.fontSize = value;
      annMgr.update(selectedPageIdx, id, updateData);
    } else if (activeTextEditor.classList.contains('pdf-text-overlay')) {
      if (formatType === 'color') activeTextEditor._origItem.textColor = value;
      if (formatType === 'fontSize') activeTextEditor._origItem.fontSize = value;
      if (formatType === 'fontFamily') activeTextEditor._origItem.fontFamily = value;
      saveInlineEdit(activeTextEditor);
    }
  }
}

function syncToolbarWithSelection() {
  if (!activeTextEditor) return;
  
  // Use document.queryCommandState to check bold/italic at current selection/caret
  const isSelectedBold = document.queryCommandState('bold');
  const isSelectedItalic = document.queryCommandState('italic');
  
  $('btn-bold').classList.toggle('active', isSelectedBold);
  $('btn-italic').classList.toggle('active', isSelectedItalic);
  
  // Color
  try {
    const selectedColor = document.queryCommandValue('foreColor');
    if (selectedColor) {
      $('color-picker').value = rgbToHex(selectedColor);
      currentColor = $('color-picker').value;
    }
  } catch (e) {}

  // Font family
  try {
    const selectedFont = document.queryCommandValue('fontName');
    if (selectedFont) {
      syncFontFamilyDropdown(selectedFont);
    }
  } catch (e) {}

  // Font size
  try {
    const selection = window.getSelection();
    if (selection.rangeCount > 0) {
      let node = selection.anchorNode;
      if (node && node.nodeType === Node.TEXT_NODE) {
        node = node.parentNode;
      }
      if (node && activeTextEditor.contains(node)) {
        let sizeFound = null;
        let fontFound = null;
        let curr = node;
        while (curr && curr !== activeTextEditor) {
          if (curr.style) {
            if (!sizeFound && curr.style.fontSize) {
              sizeFound = parseInt(curr.style.fontSize);
            }
            if (!fontFound && curr.style.fontFamily) {
              fontFound = curr.style.fontFamily;
            }
          }
          curr = curr.parentNode;
        }
        if (sizeFound) {
          $('font-size').value = sizeFound;
          fontSize = sizeFound;
        } else if (activeTextEditor.style.fontSize) {
          const sz = parseInt(activeTextEditor.style.fontSize);
          if (sz) {
            $('font-size').value = sz;
            fontSize = sz;
          }
        }
        if (fontFound) {
          syncFontFamilyDropdown(fontFound);
        }
      }
    }
  } catch (e) {}
}

document.addEventListener('selectionchange', () => {
  if (activeTextEditor) {
    const sel = window.getSelection();
    if (sel.rangeCount > 0 && !sel.isCollapsed) {
      const range = sel.getRangeAt(0);
      if (activeTextEditor.contains(range.commonAncestorContainer)) {
        saveCurrentSelection();
      }
    }
    syncToolbarWithSelection();
  }
});

// Prevent all toolbar formatting controls from stealing focus from the editor
['btn-bold', 'btn-italic'].forEach(id => {
  $(id).addEventListener('mousedown', e => e.preventDefault());
});

$('color-picker').addEventListener('input', e => {
  currentColor = e.target.value;
  applyTextFormat('color', currentColor);
});

function handleFontSizeChange(e) {
  const val = parseInt(e.target.value);
  if (!val || val < 1) return;
  fontSize = val;
  applyTextFormat('fontSize', fontSize);
}
$('font-size').addEventListener('change', handleFontSizeChange);
$('font-size').addEventListener('input', handleFontSizeChange);

$('font-family').addEventListener('change', e => {
  currentFontFamily = e.target.value;
  applyTextFormat('fontFamily', currentFontFamily);
});

$('btn-bold').addEventListener('click', () => { 
  applyTextFormat('bold', null);
});

$('btn-italic').addEventListener('click', () => { 
  applyTextFormat('italic', null);
});
$('stroke-width').addEventListener('input', e => { strokeWidth = parseInt(e.target.value); $('stroke-width-value').textContent = strokeWidth + 'px'; });

$('btn-undo').addEventListener('click', () => { const h = annMgr.undo(); if (h) refreshAnnotations(h.pageIdx); redrawAllDrawings(); });
$('btn-redo').addEventListener('click', () => { const h = annMgr.redo(); if (h) refreshAnnotations(h.pageIdx); redrawAllDrawings(); });
$('btn-delete').addEventListener('click', deleteSelected);

$('btn-back').addEventListener('click', () => {
  // Reset file input so the 'change' event fires even if the same file is selected again
  fileInput.value = '';
  imageInput.value = '';
  fileInputScan.value = '';

  // Clean up editor state
  editorMode = 'docx';
  currentFile = null;
  selectedAnn = null;
  selectedAnnEl = null;
  selectedPageIdx = null;
  activeTextEditor = null;
  drawCanvases = {};
  drawingStates = {};
  originalTextEdits = {};
  pageTextItems = {};
  pageCount = 0;
  pagesWrapper.innerHTML = '';
  thumbsContainer.innerHTML = '';
  annMgr.clear();

  // Reset tool to default
  currentTool = 'select';
  document.querySelectorAll('[data-tool]').forEach(b => b.classList.remove('active'));
  $('btn-select').classList.add('active');
  textOptions.style.display = 'flex';
  drawOptions.style.display = 'none';
  btnDelete.style.display = 'none';

  // Hide AI mode badge
  $('toolbar-ai-indicator').style.display = 'none';

  // Switch screens
  uploadScreen.style.display = 'flex';
  editorScreen.style.display = 'none';
  fileNameDisplay.style.display = 'none';
  headerActions.style.display = 'none';
});

$('btn-save').addEventListener('click', savePdf);

// ===== AI SETTINGS MODAL INITIALIZATION =====
const aiSettingsModal = $('ai-settings-modal');
const btnSettings = $('btn-settings');
const btnCloseSettings = $('btn-close-settings');
const btnSaveSettings = $('btn-save-settings');
const geminiApiKeyInput = $('gemini-api-key');
const geminiModelSelect = $('gemini-model');
const customModelGroup = $('custom-model-group');
const geminiCustomModelInput = $('gemini-custom-model');

function loadAISettings() {
  const apiKey = localStorage.getItem('gemini_api_key') || '';
  const model = localStorage.getItem('gemini_model') || 'gemini-3-flash-preview';
  const customModel = localStorage.getItem('gemini_custom_model') || '';

  geminiApiKeyInput.value = apiKey;
  geminiModelSelect.value = model;
  geminiCustomModelInput.value = customModel;

  if (model === 'custom') {
    customModelGroup.style.display = 'block';
  } else {
    customModelGroup.style.display = 'none';
  }
}

btnSettings.addEventListener('click', () => {
  loadAISettings();
  aiSettingsModal.style.display = 'flex';
});

btnCloseSettings.addEventListener('click', () => {
  aiSettingsModal.style.display = 'none';
});

geminiModelSelect.addEventListener('change', () => {
  if (geminiModelSelect.value === 'custom') {
    customModelGroup.style.display = 'block';
  } else {
    customModelGroup.style.display = 'none';
  }
});

btnSaveSettings.addEventListener('click', () => {
  const apiKey = geminiApiKeyInput.value.trim();
  const model = geminiModelSelect.value;
  const customModel = geminiCustomModelInput.value.trim();

  localStorage.setItem('gemini_api_key', apiKey);
  localStorage.setItem('gemini_model', model);
  localStorage.setItem('gemini_custom_model', customModel);

  aiSettingsModal.style.display = 'none';
  showToast('Đã lưu cấu hình AI!', 'success');
});

$('btn-zoom-in').addEventListener('click', () => setZoom(zoomLevel + 15));
$('btn-zoom-out').addEventListener('click', () => setZoom(zoomLevel - 15));
$('btn-zoom-fit').addEventListener('click', () => setZoom(100));

function setZoom(v) {
  zoomLevel = Math.max(30, Math.min(300, v));
  $('zoom-level').textContent = zoomLevel + '%';
  pagesWrapper.style.transform = `scale(${zoomLevel / 100})`;
  pagesWrapper.style.transformOrigin = 'top center';
}

document.addEventListener('keydown', e => {
  // Don't intercept when actively editing text (typing inside it)
  if (document.activeElement && document.activeElement.classList.contains('editing')) return;
  if (e.ctrlKey && e.key === 'z') { e.preventDefault(); $('btn-undo').click(); }
  if (e.ctrlKey && e.key === 'y') { e.preventDefault(); $('btn-redo').click(); }
  if (e.key === 'Delete' && selectedAnn) { e.preventDefault(); deleteSelected(); }

  // Arrow keys movement support
  if (selectedAnn && selectedAnnEl && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
    e.preventDefault();
    
    // Find page wrapper and dimensions
    const wrapper = pagesWrapper.children[selectedPageIdx];
    if (!wrapper) return;
    const pdfCanvas = wrapper.querySelector('.pdf-render');
    const pw = pdfCanvas.width, ph = pdfCanvas.height;

    // Movement step: 1 pixel by default, 5 pixels if Shift is held
    const step = e.shiftKey ? 5 : 1;
    let dx = 0;
    let dy = 0;

    if (e.key === 'ArrowUp') dy = -step;
    if (e.key === 'ArrowDown') dy = step;
    if (e.key === 'ArrowLeft') dx = -step;
    if (e.key === 'ArrowRight') dx = step;

    // Calculate new coordinates in canvas units
    let newX = selectedAnn.x + dx;
    let newY = selectedAnn.y + dy;

    // Update annotation model
    selectedAnn.x = newX;
    selectedAnn.y = newY;
    annMgr.update(selectedPageIdx, selectedAnn.id, { x: newX, y: newY });

    // Update DOM position instantly
    selectedAnnEl.style.left = (newX / pw * 100) + '%';
    selectedAnnEl.style.top = (newY / ph * 100) + '%';
  }
});

function setupDrawCanvas(canvas, pageIdx) {
  let drawing = false;
  canvas.addEventListener('mousedown', e => {
    if (currentTool !== 'draw' && currentTool !== 'highlight' && currentTool !== 'rect' && currentTool !== 'text' && currentTool !== 'eraser' && currentTool !== 'cut' && currentTool !== 'erase-stamp') return;
    drawing = true;
    const r = canvas.getBoundingClientRect();
    const sx = canvas.width / r.width, sy = canvas.height / r.height;
    const x = (e.clientX - r.left) * sx, y = (e.clientY - r.top) * sy;
    const state = drawingStates[pageIdx];
    if (currentTool === 'rect') {
      state.currentPath = { tool: 'rect', color: currentColor, width: strokeWidth, startX: x, startY: y, endX: x, endY: y };
    } else if (currentTool === 'eraser') {
      state.currentPath = { tool: 'eraser', color: currentColor, fillColor: currentColor, startX: x, startY: y, endX: x, endY: y };
    } else if (currentTool === 'cut') {
      state.currentPath = { tool: 'cut-box', startX: x, startY: y, endX: x, endY: y };
    } else if (currentTool === 'erase-stamp') {
      state.currentPath = { tool: 'erase-stamp-box', startX: x, startY: y, endX: x, endY: y };
    } else if (currentTool === 'text') {
      state.currentPath = { tool: 'text-box', color: '#6c63ff', width: 1.5, startX: x, startY: y, endX: x, endY: y };
    } else {
      state.currentPath = {
        tool: currentTool, color: currentColor,
        width: currentTool === 'highlight' ? strokeWidth * 3 : strokeWidth,
        alpha: currentTool === 'highlight' ? 0.3 : 1,
        points: [{ x, y }]
      };
    }
  });
  canvas.addEventListener('mousemove', e => {
    if (!drawing) return;
    const r = canvas.getBoundingClientRect();
    const sx = canvas.width / r.width, sy = canvas.height / r.height;
    const x = (e.clientX - r.left) * sx, y = (e.clientY - r.top) * sy;
    const state = drawingStates[pageIdx];
    if (state.currentPath.tool === 'rect' || state.currentPath.tool === 'text-box' || state.currentPath.tool === 'eraser' || state.currentPath.tool === 'cut-box' || state.currentPath.tool === 'erase-stamp-box') { 
      state.currentPath.endX = x; 
      state.currentPath.endY = y; 
    } else { 
      state.currentPath.points.push({ x, y }); 
    }
    redrawPage(pageIdx);
  });
  canvas.addEventListener('mouseup', () => {
    if (!drawing) return;
    drawing = false;
    const state = drawingStates[pageIdx];
    if (state.currentPath) {
      if (state.currentPath.tool === 'text-box') {
        const p = state.currentPath;
        const x = Math.min(p.startX, p.endX);
        const y = Math.min(p.startY, p.endY);
        const width = Math.abs(p.endX - p.startX);
        const height = Math.abs(p.endY - p.startY);
        
        state.currentPath = null;
        redrawPage(pageIdx);
        
        // If it's extremely small (like a single click), we set a default box size
        const finalW = width > 10 ? width : 150;
        const finalH = height > 10 ? height : 30;
        
        const wrapper = pagesWrapper.children[pageIdx];
        const layer = wrapper.querySelector('.text-annotations');
        
        createNewTextAnnotationWithStyle(pageIdx, layer, x, y, finalW, finalH, canvas.width, canvas.height);
      } else if (state.currentPath.tool === 'cut-box') {
        const p = state.currentPath;
        const x = Math.min(p.startX, p.endX);
        const y = Math.min(p.startY, p.endY);
        const width = Math.abs(p.endX - p.startX);
        const height = Math.abs(p.endY - p.startY);
        
        state.currentPath = null;
        redrawPage(pageIdx);
        
        if (width > 5 && height > 5) {
          performCutAndShift(pageIdx, canvas, x, y, width, height);
        }
      } else if (state.currentPath.tool === 'erase-stamp-box') {
        const p = state.currentPath;
        const x = Math.min(p.startX, p.endX);
        const y = Math.min(p.startY, p.endY);
        const width = Math.abs(p.endX - p.startX);
        const height = Math.abs(p.endY - p.startY);
        
        state.currentPath = null;
        redrawPage(pageIdx);
        
        if (width > 5 && height > 5) {
          performEraseStamp(pageIdx, canvas, x, y, width, height);
        }
      } else {
        state.paths.push(state.currentPath);
        state.currentPath = null;
        annMgr.add(pageIdx, { type: 'drawing', pathIndex: state.paths.length - 1, path: { ...state.paths[state.paths.length - 1] } });
      }
    }
  });
  canvas.addEventListener('mouseleave', () => {
    if (drawing) {
      drawing = false;
      const state = drawingStates[pageIdx];
      if (state.currentPath) {
        if (state.currentPath.tool === 'text-box' || state.currentPath.tool === 'cut-box' || state.currentPath.tool === 'erase-stamp-box') {
          state.currentPath = null;
          redrawPage(pageIdx);
        } else {
          state.paths.push(state.currentPath); 
          state.currentPath = null; 
        }
      }
    }
  });
}

function performCutAndShift(pageIdx, canvas, x, y, w, h) {
  const wrapper = canvas.parentElement;
  const pdfCanvas = wrapper.querySelector('.pdf-render');
  if (!pdfCanvas) return;

  // Round coordinates to integers to prevent sub-pixel anti-aliasing and transparent edge seams
  const rx = Math.round(x);
  const ry = Math.round(y);
  const rw = Math.round(w);
  const rh = Math.round(h);

  // Create temporary canvas to capture cropped image
  const cropCanvas = document.createElement('canvas');
  cropCanvas.width = rw;
  cropCanvas.height = rh;
  const cropCtx = cropCanvas.getContext('2d');
  
  // Fill cropCanvas with solid white first to prevent transparent black border bleeding at the edges
  cropCtx.fillStyle = '#ffffff';
  cropCtx.fillRect(0, 0, rw, rh);
  
  // Draw the section from pdfCanvas onto the cropCanvas
  cropCtx.drawImage(pdfCanvas, rx, ry, rw, rh, 0, 0, rw, rh);
  const croppedDataUrl = cropCanvas.toDataURL('image/png');

  // Draw an eraser rectangle over the original section on drawing canvas
  const state = drawingStates[pageIdx];
  const eraserPath = {
    tool: 'eraser',
    color: '#ffffff',
    fillColor: '#ffffff',
    startX: rx,
    startY: ry,
    endX: rx + rw,
    endY: ry + rh
  };
  state.paths.push(eraserPath);
  annMgr.add(pageIdx, {
    type: 'drawing',
    pathIndex: state.paths.length - 1,
    path: { ...eraserPath }
  });
  redrawPage(pageIdx);

  // Add a new image annotation in annotation manager at the original position
  const layer = wrapper.querySelector('.text-annotations');
  const annotation = {
    type: 'image',
    x: rx,
    y: ry,
    w: rw,
    h: rh,
    dataUrl: croppedDataUrl
  };
  const addedAnn = annMgr.add(pageIdx, annotation);
  
  // Render the cropped image element in the DOM
  const el = renderImageElement(addedAnn, pageIdx, layer, canvas.width, canvas.height);
  
  // Auto-select the newly cropped annotation
  selectAnnotation(addedAnn, pageIdx, el);

  // Switch tool to select
  currentTool = 'select';
  document.querySelectorAll('[data-tool]').forEach(b => {
    b.classList.toggle('active', b.dataset.tool === 'select');
  });
  updateCursors();
  
  showToast('Đã cắt vùng thành công! Sử dụng công cụ Chọn để di chuyển lên.', 'success');
}

async function callGeminiVisionAPI(apiKey, model, base64Image, promptText) {
  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
  const base64Data = base64Image.split(',')[1] || base64Image;
  
  const payload = {
    contents: [
      {
        parts: [
          { text: promptText },
          {
            inlineData: {
              mimeType: "image/png",
              data: base64Data
            }
          }
        ]
      }
    ],
    generationConfig: {
      responseMimeType: "application/json"
    }
  };

  const response = await fetch(url, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });

  if (!response.ok) {
    const errText = await response.text();
    throw new Error(`Gemini API Error (${response.status}): ${errText}`);
  }

  const result = await response.json();
  if (!result.candidates || !result.candidates[0] || !result.candidates[0].content || !result.candidates[0].content.parts || !result.candidates[0].content.parts[0]) {
    throw new Error("Không nhận được phản hồi hợp lệ từ Gemini API.");
  }
  return result.candidates[0].content.parts[0].text;
}

async function performEraseStamp(pageIdx, canvas, x, y, w, h) {
  const wrapper = canvas.parentElement;
  const pdfCanvas = wrapper.querySelector('.pdf-render');
  if (!pdfCanvas) return;

  // Round coordinates to integers
  const rx = Math.round(x);
  const ry = Math.round(y);
  const rw = Math.round(w);
  const rh = Math.round(h);

  const mode = $('erase-stamp-mode').value;
  const apiKey = localStorage.getItem('gemini_api_key') || '';
  let selectedModel = localStorage.getItem('gemini_model') || 'gemini-3-flash-preview';
  if (selectedModel === 'custom') {
    selectedModel = localStorage.getItem('gemini_custom_model') || 'gemini-3-flash-preview';
  }

  // Create temporary canvas to capture cropped image
  const cropCanvas = document.createElement('canvas');
  cropCanvas.width = rw;
  cropCanvas.height = rh;
  const cropCtx = cropCanvas.getContext('2d');
  
  // Fill cropCanvas with solid white first
  cropCtx.fillStyle = '#ffffff';
  cropCtx.fillRect(0, 0, rw, rh);
  
  // Draw the section from pdfCanvas onto the cropCanvas
  cropCtx.drawImage(pdfCanvas, rx, ry, rw, rh, 0, 0, rw, rh);
  
  const originalDataUrl = cropCanvas.toDataURL('image/png');

  let finalDataUrl = '';

  // Get image data to process pixels
  const imgData = cropCtx.getImageData(0, 0, rw, rh);

  // HSL helper
  const rgbToHsl = (r, g, b) => {
    r /= 255; g /= 255; b /= 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h, s, l = (max + min) / 2;
    if (max === min) {
      h = s = 0;
    } else {
      const d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      switch (max) {
        case r: h = (g - b) / d + (g < b ? 6 : 0); break;
        case g: h = (b - r) / d + 2; break;
        case b: h = (r - g) / d + 4; break;
      }
      h /= 6;
    }
    return [h * 360, s * 100, l * 100];
  };

  const getAlgoProcessedDataUrl = () => {
    cropCtx.fillStyle = '#ffffff';
    cropCtx.fillRect(0, 0, rw, rh);
    cropCtx.drawImage(pdfCanvas, rx, ry, rw, rh, 0, 0, rw, rh);
    
    const tempImgData = cropCtx.getImageData(0, 0, rw, rh);
    const tempData = tempImgData.data;
    
    for (let i = 0; i < tempData.length; i += 4) {
      const r = tempData[i];
      const g = tempData[i+1];
      const b = tempData[i+2];
      
      const [h, s, l] = rgbToHsl(r, g, b);
      
      const isBlue = (h >= 165 && h <= 265) && s > 8 && b > g - 10;
      const isRed = ((h >= 330 || h <= 32) && s > 12 && l > 10) || 
                    (r > 115 && r > g + 20 && r > b + 20);
      const isDark = l < 38 || (r < 100 && g < 100 && b < 100);
      
      if (isBlue) {
        const maxVal = Math.max(r, g, b);
        if (maxVal < 210) {
          tempData[i] = Math.max(0, r - 15);
          tempData[i+1] = Math.max(0, g - 10);
          tempData[i+2] = Math.min(255, b + 10);
        }
      } else if (isRed) {
        if (isDark) {
          const gray = Math.round((g + b) / 2);
          const cleanGray = Math.max(0, Math.min(50, gray));
          tempData[i] = cleanGray;
          tempData[i+1] = cleanGray;
          tempData[i+2] = cleanGray;
        } else {
          tempData[i] = 255;
          tempData[i+1] = 255;
          tempData[i+2] = 255;
        }
      } else if (isDark) {
        const gray = Math.round(r * 0.299 + g * 0.587 + b * 0.114);
        const cleanGray = gray < 60 ? 0 : gray;
        tempData[i] = cleanGray;
        tempData[i+1] = cleanGray;
        tempData[i+2] = cleanGray;
      } else {
        tempData[i] = 255;
        tempData[i+1] = 255;
        tempData[i+2] = 255;
      }
    }
    cropCtx.putImageData(tempImgData, 0, 0);
    return cropCanvas.toDataURL('image/png');
  };

  const getAiSignatureOnlyDataUrl = () => {
    cropCtx.fillStyle = '#ffffff';
    cropCtx.fillRect(0, 0, rw, rh);
    cropCtx.drawImage(pdfCanvas, rx, ry, rw, rh, 0, 0, rw, rh);
    
    const tempImgData = cropCtx.getImageData(0, 0, rw, rh);
    const tempData = tempImgData.data;
    
    for (let i = 0; i < tempData.length; i += 4) {
      const r = tempData[i];
      const g = tempData[i+1];
      const b = tempData[i+2];
      
      const [h, s, l] = rgbToHsl(r, g, b);
      const isBlue = (h >= 165 && h <= 265) && s > 8 && b > g - 10;
      
      if (isBlue) {
        const maxVal = Math.max(r, g, b);
        if (maxVal < 210) {
          tempData[i] = Math.max(0, r - 20);
          tempData[i+1] = Math.max(0, g - 15);
          tempData[i+2] = Math.min(255, b + 15);
        }
      } else {
        tempData[i] = 255;
        tempData[i+1] = 255;
        tempData[i+2] = 255;
      }
    }
    cropCtx.putImageData(tempImgData, 0, 0);
    return cropCanvas.toDataURL('image/png');
  };

  if (mode === 'ai' && apiKey) {
    showLoading("AI đang phân tích và tái tạo văn bản...");
    try {
      const promptText = `
You are a document reconstruction expert. Analyze the provided image of a signature/stamp area.
Your task is to identify and extract all black/dark text lines that exist in this image (e.g. titles like "KT. TRƯỞNG PHÒNG", "PHÓ TRƯỞNG PHÒNG", names, dates, or other information).
Do not extract the blue handwritten signature itself.
For each text line, you must output:
1. "text": The exact text content of the line.
2. "y": The vertical center position of this line as a percentage of the total image height (0 to 100).
3. "x": The horizontal center position of this line as a percentage of the total image width (0 to 100).
4. "fontSize": The font size as a percentage of the total image height (e.g., if the text line height is 15% of the box, return 15). Typical values range from 8 to 22.
5. "isBold": true if the text is bold, false otherwise.
6. "fontFamily": The detected font family category ('serif', 'sans-serif', 'monospace').

Return the result STRICTLY as a JSON array of objects, for example:
[
  {"text": "KT. TRƯỞNG PHÒNG", "x": 50, "y": 20, "fontSize": 14, "isBold": true, "fontFamily": "sans-serif"},
  {"text": "PHÓ TRƯỞNG PHÒNG", "x": 50, "y": 38, "fontSize": 14, "isBold": true, "fontFamily": "sans-serif"}
]
Do not include any explanation, markdown formatting, or text outside the JSON block.
`;

      const responseText = await callGeminiVisionAPI(apiKey, selectedModel, originalDataUrl, promptText);
      
      let textItems = [];
      try {
        let cleanText = responseText.trim();
        if (cleanText.startsWith('```')) {
          cleanText = cleanText.replace(/^```json\s*/, '').replace(/```$/, '').trim();
        }
        textItems = JSON.parse(cleanText);
      } catch (err) {
        console.error("Failed to parse Gemini response:", responseText, err);
        throw new Error("Không thể phân tích định dạng JSON từ phản hồi của AI.");
      }

      const layer = wrapper.querySelector('.text-annotations');
      const pw = pdfCanvas.width, ph = pdfCanvas.height;

      if (Array.isArray(textItems)) {
        for (const item of textItems) {
          if (!item.text || !item.text.trim()) continue;
          
          const fs = Math.round((item.fontSize / 100) * rh);
          const estimatedWidth = item.text.length * fs * 0.55;
          const annX = Math.max(0, Math.min(pw, rx + (item.x / 100) * rw - estimatedWidth / 2));
          const annY = Math.max(0, Math.min(ph, ry + (item.y / 100) * rh - fs / 2));

          const annText = annMgr.add(pageIdx, {
            type: 'text',
            x: annX,
            y: annY,
            text: item.text,
            color: '#000000',
            fontSize: fs,
            bold: item.isBold !== false,
            italic: false,
            fontName: null,
            fontFamily: item.fontFamily === 'serif' ? "'Times New Roman', Times, serif" : "Arial, Helvetica, sans-serif"
          });
          renderNewTextElement(annText, pageIdx, layer, pw, ph, false);
        }
      }
      
      finalDataUrl = getAiSignatureOnlyDataUrl();
      showToast('Đã xoá mộc đỏ và tái tạo văn bản bằng AI thành công!', 'success');
    } catch (e) {
      console.error(e);
      showToast('Lỗi AI: ' + e.message + '. Đã tự động dùng thuật toán dự phòng.', 'warning');
      finalDataUrl = getAlgoProcessedDataUrl();
    } finally {
      hideLoading();
    }
  } else {
    if (mode === 'ai' && !apiKey) {
      showToast('Chưa cấu hình API Key. Đã tự động chuyển sang chế độ Thuật toán để giữ lại chữ và chữ ký sắc nét!', 'warning');
    } else {
      showToast('Đã xoá mộc đỏ và giữ lại chữ ký thành công!', 'success');
    }
    finalDataUrl = getAlgoProcessedDataUrl();
  }

  // Draw an eraser rectangle over the original section on drawing canvas
  const state = drawingStates[pageIdx];
  const eraserPath = {
    tool: 'eraser',
    color: '#ffffff',
    fillColor: '#ffffff',
    startX: rx,
    startY: ry,
    endX: rx + rw,
    endY: ry + rh
  };
  state.paths.push(eraserPath);
  annMgr.add(pageIdx, {
    type: 'drawing',
    pathIndex: state.paths.length - 1,
    path: { ...eraserPath }
  });
  redrawPage(pageIdx);

  // Add a new image annotation in annotation manager at the original position
  const layer = wrapper.querySelector('.text-annotations');
  const pw = pdfCanvas.width, ph = pdfCanvas.height;
  const annotation = {
    type: 'image',
    x: rx,
    y: ry,
    w: rw,
    h: rh,
    dataUrl: finalDataUrl
  };
  const addedAnn = annMgr.add(pageIdx, annotation);
  
  // Render the cropped image element in the DOM
  const el = renderImageElement(addedAnn, pageIdx, layer, pw, ph);

  // Auto-select the newly cropped annotation
  selectAnnotation(addedAnn, pageIdx, el);

  // Switch tool to select
  currentTool = 'select';
  document.querySelectorAll('[data-tool]').forEach(b => {
    b.classList.toggle('active', b.dataset.tool === 'select');
  });
  updateCursors();
}

function redrawPage(pageIdx, isExport = false, hideErasers = false) {
  const canvas = drawCanvases[pageIdx];
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  const state = drawingStates[pageIdx];

  // Rebuild state.paths from annMgr's drawing annotations to stay in perfect sync during Undo/Redo
  const drawings = annMgr.getPage(pageIdx).filter(a => {
    if (a.type !== 'drawing') return false;
    if (hideErasers && a.path && a.path.tool === 'eraser') return false;
    return true;
  });
  state.paths = drawings.map(a => a.path);

  const drawPath = p => {
    ctx.save();
    ctx.globalAlpha = p.alpha || 1;
    ctx.strokeStyle = p.color;
    ctx.lineWidth = p.width;
    ctx.lineCap = 'round';
    ctx.lineJoin = 'round';
    if (p.tool === 'rect') {
      ctx.strokeRect(p.startX, p.startY, p.endX - p.startX, p.endY - p.startY);
    } else if (p.tool === 'eraser') {
      ctx.fillStyle = p.fillColor || '#ffffff';
      ctx.fillRect(p.startX, p.startY, p.endX - p.startX, p.endY - p.startY);
      if (!isExport) {
        ctx.strokeStyle = 'rgba(0, 0, 0, 0.2)';
        ctx.lineWidth = 1;
        ctx.setLineDash([4, 4]);
        ctx.strokeRect(p.startX, p.startY, p.endX - p.startX, p.endY - p.startY);
        ctx.setLineDash([]);
      }
    } else if (p.tool === 'cut-box' || p.tool === 'erase-stamp-box') {
      ctx.strokeStyle = p.tool === 'erase-stamp-box' ? '#38ef7d' : '#ff4757';
      ctx.lineWidth = 1.5;
      ctx.setLineDash([4, 4]);
      ctx.strokeRect(p.startX, p.startY, p.endX - p.startX, p.endY - p.startY);
      ctx.setLineDash([]);
    } else if (p.tool === 'text-box') {
      ctx.setLineDash([5, 5]);
      ctx.strokeRect(p.startX, p.startY, p.endX - p.startX, p.endY - p.startY);
      ctx.setLineDash([]);
    } else {
      ctx.beginPath();
      if (p.points && p.points.length > 0) {
        ctx.moveTo(p.points[0].x, p.points[0].y);
        for (let i = 1; i < p.points.length; i++) ctx.lineTo(p.points[i].x, p.points[i].y);
      }
      ctx.stroke();
    }
    ctx.restore();
  };
  state.paths.forEach(drawPath);
  if (state.currentPath) drawPath(state.currentPath);
}

function redrawAllDrawings() {
  Object.keys(drawCanvases).forEach(idx => redrawPage(parseInt(idx)));
}

// ===== IMAGE =====
imageInput.addEventListener('change', e => {
  const file = e.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = ev => {
    const pageIdx = 0;
    const wrapper = pagesWrapper.children[pageIdx];
    const pdfCanvas = wrapper.querySelector('.pdf-render');
    const layer = wrapper.querySelector('.text-annotations');
    const pw = pdfCanvas.width, ph = pdfCanvas.height;
    const ann = annMgr.add(pageIdx, {
      type: 'image', x: pw * 0.1, y: ph * 0.1, w: pw * 0.3, h: ph * 0.3, dataUrl: ev.target.result
    });
    renderImageElement(ann, pageIdx, layer, pw, ph);
    showToast('Đã thêm ảnh', 'success');
  };
  reader.readAsDataURL(file);
  imageInput.value = '';
  document.querySelector('[data-tool="select"]').click();
});

function renderImageElement(ann, pageIdx, layer, pw, ph) {
  const el = document.createElement('div');
  el.className = 'image-annotation';
  el.dataset.id = ann.id;
  el.style.left = (ann.x / pw * 100) + '%';
  el.style.top = (ann.y / ph * 100) + '%';
  el.style.width = (ann.w / pw * 100) + '%';
  el.style.height = (ann.h / ph * 100) + '%';
  const img = document.createElement('img');
  img.src = ann.dataUrl;
  el.appendChild(img);
  const handle = document.createElement('div');
  handle.className = 'resize-handle';
  el.appendChild(handle);

  let dragging = false, dx, dy;
  el.addEventListener('mousedown', e => {
    if (e.target === handle || currentTool !== 'select') return;
    dragging = true;
    dx = e.clientX - el.getBoundingClientRect().left;
    dy = e.clientY - el.getBoundingClientRect().top;
    selectAnnotation(ann, pageIdx, el);
    e.preventDefault();
  });
  document.addEventListener('mousemove', e => {
    if (!dragging) return;
    const rect = layer.getBoundingClientRect();
    el.style.left = ((e.clientX - rect.left - dx) / rect.width * 100) + '%';
    el.style.top = ((e.clientY - rect.top - dy) / rect.height * 100) + '%';
  });
  document.addEventListener('mouseup', () => {
    if (dragging) { dragging = false; annMgr.update(pageIdx, ann.id, { x: parseFloat(el.style.left) / 100 * pw, y: parseFloat(el.style.top) / 100 * ph }); }
  });

  let resizing = false, startW, startH, startX2, startY2;
  handle.addEventListener('mousedown', e => {
    resizing = true; startW = el.offsetWidth; startH = el.offsetHeight;
    startX2 = e.clientX; startY2 = e.clientY; e.stopPropagation(); e.preventDefault();
  });
  document.addEventListener('mousemove', e => {
    if (!resizing) return;
    const rect = layer.getBoundingClientRect();
    el.style.width = (Math.max(30, startW + (e.clientX - startX2)) / rect.width * 100) + '%';
    el.style.height = (Math.max(30, startH + (e.clientY - startY2)) / rect.height * 100) + '%';
  });
  document.addEventListener('mouseup', () => {
    if (resizing) { resizing = false; annMgr.update(pageIdx, ann.id, { w: parseFloat(el.style.width) / 100 * pw, h: parseFloat(el.style.height) / 100 * ph }); }
  });

  el.addEventListener('click', e => {
    if (currentTool === 'select') { selectAnnotation(ann, pageIdx, el); e.stopPropagation(); }
  });
  layer.appendChild(el);
  return el;
}

// ===== DOM-TO-CANVAS TEXT RENDERING (pixel-perfect match with preview) =====
/**
 * Renders text using HTML/CSS (identical to preview) and captures it as an Image.
 * Uses an offscreen canvas with the actual DOM rendering engine instead of canvas fillText,
 * which tends to render text slightly bolder/darker.
 */
const inlinedFontCache = new Map();

async function inlineFontRules(fontFaceCSS) {
  const blobRegex = /url\((["']?)(blob:[^"')]+)\1\)/g;
  let match;
  let css = fontFaceCSS;
  const promises = [];
  
  const blobUrls = new Set();
  blobRegex.lastIndex = 0;
  while ((match = blobRegex.exec(fontFaceCSS)) !== null) {
    blobUrls.add(match[2]);
  }
  
  for (const url of blobUrls) {
    if (inlinedFontCache.has(url)) {
      continue;
    }
    promises.push((async () => {
      try {
        const response = await fetch(url);
        const blob = await response.blob();
        return new Promise((resolve) => {
          const reader = new FileReader();
          reader.onloadend = () => {
            inlinedFontCache.set(url, reader.result);
            resolve();
          };
          reader.readAsDataURL(blob);
        });
      } catch (err) {
        console.warn('Failed to inline font blob URL:', url, err);
        inlinedFontCache.set(url, null);
      }
    })());
  }
  
  await Promise.all(promises);
  
  css = css.replace(blobRegex, (fullMatch, quote, url) => {
    const dataUrl = inlinedFontCache.get(url);
    if (dataUrl) {
      return `url("${dataUrl}")`;
    }
    return 'none';
  });
  
  return css;
}

async function renderTextViaDOM(text, fontFamily, fontSize, color, bold, italic, width, height, isHTML = false) {
  return new Promise((resolve, reject) => {
    const cleanedText = text
      .replace(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\xA0\u2000-\u200A\u202F\u205F\u3000\uF020]/g, ' ')
      .replace(/[\xAD\u200B-\u200F\uFEFF]/g, '');

    // Create a temporary container matching the inline editor styling
    const container = document.createElement('div');
    container.style.cssText = `
      position: fixed; left: -9999px; top: -9999px;
      font-family: ${fontFamily};
      font-size: ${fontSize}px;
      font-weight: ${bold ? 'bold' : 'normal'};
      font-style: ${italic ? 'italic' : 'normal'};
      color: ${color};
      line-height: 1.35;
      white-space: pre-wrap;
      letter-spacing: 0;
      padding: 0;
      margin: 0;
      background: transparent;
      -webkit-font-smoothing: antialiased;
    `;
    if (isHTML) {
      container.innerHTML = cleanedText;
    } else {
      container.textContent = cleanedText;
    }
    document.body.appendChild(container);

    // Measure actual rendered size
    const rect = container.getBoundingClientRect();
    const w = Math.ceil(Math.max(rect.width, width) + 4);
    const h = Math.ceil(Math.max(rect.height, height) + 4);

    // Collect all @font-face rules from the document (PDF.js injects custom fonts)
    let fontFaceCSS = '';
    try {
      for (const sheet of document.styleSheets) {
        try {
          for (const rule of sheet.cssRules || []) {
            if (rule instanceof CSSFontFaceRule) {
              fontFaceCSS += rule.cssText + '\n';
            }
          }
        } catch (e) { /* cross-origin stylesheet, skip */ }
      }
    } catch (e) { /* ignore */ }

    inlineFontRules(fontFaceCSS).then(inlinedCSS => {
      // Change position styles so it renders correctly inside the SVG frame
      container.style.position = 'relative';
      container.style.left = '0';
      container.style.top = '0';

      // Build SVG foreignObject with embedded HTML
      const htmlContent = container.outerHTML;
      document.body.removeChild(container);

      const svgNS = 'http://www.w3.org/2000/svg';
      const svgStr = `<svg xmlns="${svgNS}" width="${w}" height="${h}">
        <defs><style>${inlinedCSS}</style></defs>
        <foreignObject width="100%" height="100%">
          <div xmlns="http://www.w3.org/1999/xhtml" style="width:${w}px;height:${h}px;overflow:hidden;">
            ${htmlContent}
          </div>
        </foreignObject>
      </svg>`;

      const svgBlob = new Blob([svgStr], { type: 'image/svg+xml;charset=utf-8' });
      const url = URL.createObjectURL(svgBlob);

      const img = new Image();
      const timeoutId = setTimeout(() => {
        URL.revokeObjectURL(url);
        reject(new Error('SVG rendering timeout'));
      }, 5000);

      img.onload = () => {
        clearTimeout(timeoutId);
        try {
          const offCanvas = document.createElement('canvas');
          offCanvas.width = w;
          offCanvas.height = h;
          const offCtx = offCanvas.getContext('2d');
          offCtx.drawImage(img, 0, 0);
          URL.revokeObjectURL(url);

          const testData = offCtx.getImageData(0, 0, w, Math.min(h, 10)).data;
          let hasContent = false;
          for (let i = 3; i < testData.length; i += 4) {
            if (testData[i] > 0) { hasContent = true; break; }
          }

          if (hasContent) {
            resolve(offCanvas);
          } else {
            reject(new Error('SVG foreignObject rendered blank'));
          }
        } catch (err) {
          URL.revokeObjectURL(url);
          reject(err);
        }
      };
      img.onerror = () => {
        clearTimeout(timeoutId);
        URL.revokeObjectURL(url);
        reject(new Error('SVG image load failed'));
      };
      img.src = url;
    }).catch(reject);
  });
}

// ===== SAVE PDF =====
async function savePdf() {
  if (!currentFile) return;
  showLoading('Đang lưu PDF...');
  try {
    const drawDataUrls = {};
    
    for (let idx = 0; idx < pageCount; idx++) {
      const canvas = drawCanvases[idx];
      if (!canvas) continue;

      // Redraw without helper borders and without erasers for export (we draw erasers natively in PDF instead)
      redrawPage(idx, true, true);

      // Create a temporary high-resolution canvas to merge drawings, text annotations, and text edits
      const mergedCanvas = document.createElement('canvas');
      mergedCanvas.width = canvas.width;
      mergedCanvas.height = canvas.height;
      // Use geometricPrecision to get thinner/sharper text more similar to HTML CSS rendering
      mergedCanvas.style.textRendering = 'geometricPrecision';
      const ctx = mergedCanvas.getContext('2d');

      // 1. Draw free hand drawings, highlights, and rectangles from original drawCanvas
      ctx.drawImage(canvas, 0, 0);

      // 2. Render all original text edits using DOM-to-Canvas capture (SVG foreignObject)
      // This produces identical output to the HTML preview, avoiding Canvas 2D fillText
      // which renders text slightly bolder/darker than CSS text rendering.
      const textEdits = originalTextEdits[idx] || [];
      for (const edit of textEdits) {
        if (edit.newText && edit.newText.trim()) {
          const item = edit.origItem;
          const fontName = item.fontName || '';
          const fnLower = fontName.toLowerCase();
          const isBold = edit.isBold !== undefined ? edit.isBold : (fnLower.includes('bold') || fnLower.includes('gothic') || fnLower.includes('heavy') || fnLower.includes('black'));
          const isItalic = edit.isItalic !== undefined ? edit.isItalic : (fnLower.includes('italic') || fnLower.includes('oblique'));
          const drawBold = isBold;
          const drawItalic = isItalic;
          const textColor = edit.textColor || item.textColor || '#000000';
          const fontFamily = item.fontFamily || 'Arial, Helvetica, sans-serif';

          // Try DOM-to-Canvas via SVG foreignObject for pixel-perfect match with preview
          try {
            const textImg = await renderTextViaDOM(
              edit.newHtml || edit.newText, fontFamily, item.fontSize, textColor,
              drawBold, drawItalic, item.width, item.height, !!edit.newHtml
            );
            ctx.drawImage(textImg, item.x, item.y);
          } catch (domErr) {
            // Fallback to canvas fillText if DOM rendering fails
            console.warn('DOM text render failed, using fillText fallback', domErr);
            const fontParts = [];
            if (drawItalic) fontParts.push('italic');
            if (drawBold) fontParts.push('bold');
            fontParts.push(`${item.fontSize}px`);
            fontParts.push(fontFamily);

            ctx.save();
            ctx.font = fontParts.join(' ');
            ctx.fillStyle = textColor;
            ctx.textBaseline = 'top';
            const lines = edit.newText.split('\n');
            const lineHeight = item.fontSize * 1.35;
            lines.forEach((line, lIdx) => {
              ctx.fillText(line, item.x, item.y + lIdx * lineHeight);
            });
            ctx.restore();
          }
        }
      }

      // 3. Render all user-added text annotations onto this canvas
      const pageAnns = annMgr.getAll()[idx] || [];
      for (const ann of pageAnns) {
        if (ann.type === 'text' && ann.text && ann.text.trim()) {
          const fontName = ann.fontName || 'Arial';
          // User-created annotations always use standard system fonts,
          // so bold/italic should always be applied directly.
          const drawBold = !!ann.bold;
          const drawItalic = !!ann.italic;
          const annW = ann.w || 9999;
          const annH = ann.h || ann.fontSize * 2;
          const fontFamily = ann.fontFamily || 'Arial, Helvetica, sans-serif';

          try {
            const textImg = await renderTextViaDOM(
              ann.html || ann.text, fontFamily, ann.fontSize, ann.color || '#000000',
              drawBold, drawItalic, annW, annH, !!ann.html
            );
            ctx.drawImage(textImg, ann.x, ann.y);
          } catch (domErr) {
            // Fallback to canvas fillText
            ctx.save();
            const fontParts = [];
            if (drawItalic) fontParts.push('italic');
            if (drawBold) fontParts.push('bold');
            fontParts.push(`${ann.fontSize}px`);
            fontParts.push(fontFamily);

            ctx.font = fontParts.join(' ');
            ctx.fillStyle = ann.color || '#000000';
            ctx.textBaseline = 'top';
            
            const wrapText = (text, maxWidth) => {
              const words = text.split(' ');
              const wrappedLines = [];
              let currentLine = words[0] || '';

              for (let i = 1; i < words.length; i++) {
                const word = words[i];
                const testWidth = ctx.measureText(currentLine + ' ' + word).width;
                if (testWidth < maxWidth) {
                  currentLine += ' ' + word;
                } else {
                  wrappedLines.push(currentLine);
                  currentLine = word;
                }
              }
              wrappedLines.push(currentLine);
              return wrappedLines;
            };

            const rawLines = ann.text.split('\n');
            const lines = [];
            
            if (ann.w) {
              rawLines.forEach(rl => {
                const wrapped = wrapText(rl, ann.w);
                lines.push(...wrapped);
              });
            } else {
              lines.push(...rawLines);
            }

            const lineHeight = ann.fontSize * 1.2;
            lines.forEach((line, lIdx) => {
              ctx.fillText(line, ann.x, ann.y + lIdx * lineHeight);
            });
            ctx.restore();
          }
        }
      }

      // Export the merged layer to PNG data URL
      drawDataUrls[idx] = mergedCanvas.toDataURL('image/png');

      // Restore helper borders for editor view
      redrawPage(idx, false);
    }

    await PdfSaver.save(currentFile, annMgr.getAll(), drawDataUrls, renderer.scale, originalTextEdits);
    showToast('Đã lưu PDF thành công!', 'success');
  } catch (e) {
    console.error(e);
    showToast('Lỗi lưu PDF: ' + e.message, 'error');
  }
  hideLoading();
}

// Deselect
document.addEventListener('click', e => {
  if (!e.target.closest('.text-annotation') && !e.target.closest('.image-annotation') && !e.target.closest('.pdf-text-overlay')) {
    document.querySelectorAll('.text-annotation.selected, .image-annotation.selected').forEach(el => el.classList.remove('selected'));
    selectedAnn = null;
    btnDelete.style.display = 'none';
  }
});

updateCursors();
